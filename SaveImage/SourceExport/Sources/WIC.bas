Attribute VB_Name = "WIC"
' //
' // Module WIC.bas - Helpes functions for Windows Imaging Component
' // By The trick 2018 (c)
' //

Option Explicit

[IgnoreWarnings(TB0015)]
Private Declare PtrSafe Function GetMem8 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
                                              
Public Function WICRect( _
                Optional ByVal lX As Long, _
                Optional ByVal lY As Long, _
                Optional ByVal lWidth As Long, _
                Optional ByVal lHeight As Long) As WICRect
    
    With WICRect
    
    .x = lX
    .y = lY
    .width = lWidth
    .Height = lHeight
    
    End With
    
End Function

Public Function CLSID_WICImagingFactory() As UUID
    GetMem8 505010466981379.7474@, CLSID_WICImagingFactory
    GetMem8 74220797839377.9105@, ByVal VarPtr(CLSID_WICImagingFactory) + 8
End Function

Public Function GUID_VendorMicrosoft() As UUID
    GetMem8 501079767522104.3658@, GUID_VendorMicrosoft
    GetMem8 311041546204258.1671@, ByVal VarPtr(GUID_VendorMicrosoft) + 8
End Function

Public Function GUID_VendorMicrosoftBuiltIn() As UUID
    GetMem8 505614238603609.3181@, GUID_VendorMicrosoftBuiltIn
    GetMem8 373954245155432.9774@, ByVal VarPtr(GUID_VendorMicrosoftBuiltIn) + 8
End Function

Public Function CLSID_WICBmpDecoder() As UUID
    GetMem8 461548235470498.2114@, CLSID_WICBmpDecoder
    GetMem8 865790619999756.9951@, ByVal VarPtr(CLSID_WICBmpDecoder) + 8
End Function

Public Function CLSID_WICPngDecoder() As UUID
    GetMem8 553895306898850.6491@, CLSID_WICPngDecoder
    GetMem8 589280763143087.3014@, ByVal VarPtr(CLSID_WICPngDecoder) + 8
End Function

Public Function CLSID_WICIcoDecoder() As UUID
    GetMem8 538100777506072.0863@, CLSID_WICIcoDecoder
    GetMem8 -8615367918846.9848@, ByVal VarPtr(CLSID_WICIcoDecoder) + 8
End Function

Public Function CLSID_WICJpegDecoder() As UUID
    GetMem8 489397963127826.9568@, CLSID_WICJpegDecoder
    GetMem8 -384116409588072.765@, ByVal VarPtr(CLSID_WICJpegDecoder) + 8
End Function

Public Function CLSID_WICGifDecoder() As UUID
    GetMem8 520295599500255.9036@, CLSID_WICGifDecoder
    GetMem8 -473244211377869.6542@, ByVal VarPtr(CLSID_WICGifDecoder) + 8
End Function

Public Function CLSID_WICTiffDecoder() As UUID
    GetMem8 530523831538486.6265@, CLSID_WICTiffDecoder
    GetMem8 313143072734960.2443@, ByVal VarPtr(CLSID_WICTiffDecoder) + 8
End Function

Public Function CLSID_WICWmpDecoder() As UUID
    GetMem8 528276117495512.5814@, CLSID_WICWmpDecoder
    GetMem8 94516622399445.9822@, ByVal VarPtr(CLSID_WICWmpDecoder) + 8
End Function

Public Function CLSID_WICBmpEncoder() As UUID
    GetMem8 517261993744888.9268@, CLSID_WICBmpEncoder
    GetMem8 -906370146900237.657@, ByVal VarPtr(CLSID_WICBmpEncoder) + 8
End Function

Public Function CLSID_WICPngEncoder() As UUID
    GetMem8 474440962245844.0041@, CLSID_WICPngEncoder
    GetMem8 -254785275739725.63@, ByVal VarPtr(CLSID_WICPngEncoder) + 8
End Function

Public Function CLSID_WICJpegEncoder() As UUID
    GetMem8 510603782837849.0305@, CLSID_WICJpegEncoder
    GetMem8 854977537333679.0198@, ByVal VarPtr(CLSID_WICJpegEncoder) + 8
End Function

Public Function CLSID_WICGifEncoder() As UUID
    GetMem8 465673425564829.8392@, CLSID_WICGifEncoder
    GetMem8 -477899909690971.7114@, ByVal VarPtr(CLSID_WICGifEncoder) + 8
End Function

Public Function CLSID_WICTiffEncoder() As UUID
    GetMem8 550315245835712.872@, CLSID_WICTiffEncoder
    GetMem8 -170778897152706.5431@, ByVal VarPtr(CLSID_WICTiffEncoder) + 8
End Function

Public Function CLSID_WICWmpEncoder() As UUID
    GetMem8 495786698674044.2059@, CLSID_WICWmpEncoder
    GetMem8 -442300938893432.691@, ByVal VarPtr(CLSID_WICWmpEncoder) + 8
End Function

Public Function GUID_ContainerFormatBmp() As UUID
    GetMem8 472230237733347.955@, GUID_ContainerFormatBmp
    GetMem8 -203240613009005.4723@, ByVal VarPtr(GUID_ContainerFormatBmp) + 8
End Function

Public Function GUID_ContainerFormatPng() As UUID
    GetMem8 513310219115357.6692@, GUID_ContainerFormatPng
    GetMem8 -578758373312287.1877@, ByVal VarPtr(GUID_ContainerFormatPng) + 8
End Function

Public Function GUID_ContainerFormatIco() As UUID
    GetMem8 548290776336592.9156@, GUID_ContainerFormatIco
    GetMem8 241826005721780.0849@, ByVal VarPtr(GUID_ContainerFormatIco) + 8
End Function

Public Function GUID_ContainerFormatJpeg() As UUID
    GetMem8 574809547874950.4938@, GUID_ContainerFormatJpeg
    GetMem8 627367042164613.136@, ByVal VarPtr(GUID_ContainerFormatJpeg) + 8
End Function

Public Function GUID_ContainerFormatTiff() As UUID
    GetMem8 569589564446839.9152@, GUID_ContainerFormatTiff
    GetMem8 -666287334752025.8666@, ByVal VarPtr(GUID_ContainerFormatTiff) + 8
End Function

Public Function GUID_ContainerFormatGif() As UUID
    GetMem8 552971368767595.0593@, GUID_ContainerFormatGif
    GetMem8 -650490558910224.7268@, ByVal VarPtr(GUID_ContainerFormatGif) + 8
End Function

Public Function GUID_ContainerFormatWmp() As UUID
    GetMem8 499004828621075.1658@, GUID_ContainerFormatWmp
    GetMem8 542065584542065.7553@, ByVal VarPtr(GUID_ContainerFormatWmp) + 8
End Function

Public Function CLSID_WICImagingCategories() As UUID
    GetMem8 505416319137715.4944@, CLSID_WICImagingCategories
    GetMem8 -910006832878856.05@, ByVal VarPtr(CLSID_WICImagingCategories) + 8
End Function

Public Function CATID_WICBitmapDecoders() As UUID
    GetMem8 519337927997609.7847@, CATID_WICBitmapDecoders
    GetMem8 -320710703730228.795@, ByVal VarPtr(CATID_WICBitmapDecoders) + 8
End Function

Public Function CATID_WICBitmapEncoders() As UUID
    GetMem8 562533583260099.855@, CATID_WICBitmapEncoders
    GetMem8 911264890302663.132@, ByVal VarPtr(CATID_WICBitmapEncoders) + 8
End Function

Public Function CATID_WICPixelFormats() As UUID
    GetMem8 513376674311824.3599@, CATID_WICPixelFormats
    GetMem8 80885593766290.8041@, ByVal VarPtr(CATID_WICPixelFormats) + 8
End Function

Public Function CATID_WICFormatConverters() As UUID
    GetMem8 531924272953831.7032@, CATID_WICFormatConverters
    GetMem8 519785243572566.7987@, ByVal VarPtr(CATID_WICFormatConverters) + 8
End Function

Public Function CATID_WICMetadataReader() As UUID
    GetMem8 553561163511729.8904@, CATID_WICMetadataReader
    GetMem8 -512395429199575.1746@, ByVal VarPtr(CATID_WICMetadataReader) + 8
End Function

Public Function CATID_WICMetadataWriter() As UUID
    GetMem8 544686349601287.21@, CATID_WICMetadataWriter
    GetMem8 332438545091324.1789@, ByVal VarPtr(CATID_WICMetadataWriter) + 8
End Function

Public Function CLSID_WICDefaultFormatConverter() As UUID
    GetMem8 541099257525325.462@, CLSID_WICDefaultFormatConverter
    GetMem8 -105772104052366.5524@, ByVal VarPtr(CLSID_WICDefaultFormatConverter) + 8
End Function

Public Function CLSID_WICFormatConverterHighColor() As UUID
    GetMem8 525812762642047.4964@, CLSID_WICFormatConverterHighColor
    GetMem8 125214773933391.5321@, ByVal VarPtr(CLSID_WICFormatConverterHighColor) + 8
End Function

Public Function CLSID_WICFormatConverterNChannel() As UUID
    GetMem8 517684009647660.5362@, CLSID_WICFormatConverterNChannel
    GetMem8 -102091503791535.9323@, ByVal VarPtr(CLSID_WICFormatConverterNChannel) + 8
End Function

Public Function CLSID_WICFormatConverterWMPhoto() As UUID
    GetMem8 509662122644059.5243@, CLSID_WICFormatConverterWMPhoto
    GetMem8 -281018185149373.0389@, ByVal VarPtr(CLSID_WICFormatConverterWMPhoto) + 8
End Function

Public Function GUID_WICPixelFormatDontCare() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormatDontCare
    GetMem8 5673201026501.9825@, ByVal VarPtr(GUID_WICPixelFormatDontCare) + 8
End Function

Public Function GUID_WICPixelFormat1bppIndexed() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat1bppIndexed
    GetMem8 12878960430294.7761@, ByVal VarPtr(GUID_WICPixelFormat1bppIndexed) + 8
End Function

Public Function GUID_WICPixelFormat2bppIndexed() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat2bppIndexed
    GetMem8 20084719834087.5697@, ByVal VarPtr(GUID_WICPixelFormat2bppIndexed) + 8
End Function

Public Function GUID_WICPixelFormat4bppIndexed() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat4bppIndexed
    GetMem8 27290479237880.3633@, ByVal VarPtr(GUID_WICPixelFormat4bppIndexed) + 8
End Function

Public Function GUID_WICPixelFormat8bppIndexed() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat8bppIndexed
    GetMem8 34496238641673.1569@, ByVal VarPtr(GUID_WICPixelFormat8bppIndexed) + 8
End Function

Public Function GUID_WICPixelFormatBlackWhite() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormatBlackWhite
    GetMem8 41701998045465.9505@, ByVal VarPtr(GUID_WICPixelFormatBlackWhite) + 8
End Function

Public Function GUID_WICPixelFormat2bppGray() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat2bppGray
    GetMem8 48907757449258.7441@, ByVal VarPtr(GUID_WICPixelFormat2bppGray) + 8
End Function

Public Function GUID_WICPixelFormat4bppGray() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat4bppGray
    GetMem8 56113516853051.5377@, ByVal VarPtr(GUID_WICPixelFormat4bppGray) + 8
End Function

Public Function GUID_WICPixelFormat8bppGray() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat8bppGray
    GetMem8 63319276256844.3313@, ByVal VarPtr(GUID_WICPixelFormat8bppGray) + 8
End Function

Public Function GUID_WICPixelFormat8bppAlpha() As UUID
    GetMem8 471130917170977.2054@, GUID_WICPixelFormat8bppAlpha
    GetMem8 -766267726677937.2118@, ByVal VarPtr(GUID_WICPixelFormat8bppAlpha) + 8
End Function

Public Function GUID_WICPixelFormat16bppBGR555() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat16bppBGR555
    GetMem8 70525035660637.1249@, ByVal VarPtr(GUID_WICPixelFormat16bppBGR555) + 8
End Function

Public Function GUID_WICPixelFormat16bppBGR565() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat16bppBGR565
    GetMem8 77730795064429.9185@, ByVal VarPtr(GUID_WICPixelFormat16bppBGR565) + 8
End Function

Public Function GUID_WICPixelFormat16bppBGRA5551() As UUID
    GetMem8 528777340775382.9419@, GUID_WICPixelFormat16bppBGRA5551
    GetMem8 -327663865128437.1795@, ByVal VarPtr(GUID_WICPixelFormat16bppBGRA5551) + 8
End Function

Public Function GUID_WICPixelFormat16bppGray() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat16bppGray
    GetMem8 84936554468222.7121@, ByVal VarPtr(GUID_WICPixelFormat16bppGray) + 8
End Function

Public Function GUID_WICPixelFormat24bppBGR() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat24bppBGR
    GetMem8 92142313872015.5057@, ByVal VarPtr(GUID_WICPixelFormat24bppBGR) + 8
End Function

Public Function GUID_WICPixelFormat24bppRGB() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat24bppRGB
    GetMem8 99348073275808.2993@, ByVal VarPtr(GUID_WICPixelFormat24bppRGB) + 8
End Function

Public Function GUID_WICPixelFormat32bppBGR() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat32bppBGR
    GetMem8 106553832679601.0929@, ByVal VarPtr(GUID_WICPixelFormat32bppBGR) + 8
End Function

Public Function GUID_WICPixelFormat32bppBGRA() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat32bppBGRA
    GetMem8 113759592083393.8865@, ByVal VarPtr(GUID_WICPixelFormat32bppBGRA) + 8
End Function

Public Function GUID_WICPixelFormat32bppPBGRA() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat32bppPBGRA
    GetMem8 120965351487186.6801@, ByVal VarPtr(GUID_WICPixelFormat32bppPBGRA) + 8
End Function

Public Function GUID_WICPixelFormat32bppGrayFloat() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat32bppGrayFloat
    GetMem8 128171110890979.4737@, ByVal VarPtr(GUID_WICPixelFormat32bppGrayFloat) + 8
End Function

Public Function GUID_WICPixelFormat32bppRGBA() As UUID
    GetMem8 489018192834066.3597@, GUID_WICPixelFormat32bppRGBA
    GetMem8 -164996430182516.9241@, ByVal VarPtr(GUID_WICPixelFormat32bppRGBA) + 8
End Function

Public Function GUID_WICPixelFormat32bppPRGBA() As UUID
    GetMem8 556409745258136.5328@, GUID_WICPixelFormat32bppPRGBA
    GetMem8 -497706277213299.7463@, ByVal VarPtr(GUID_WICPixelFormat32bppPRGBA) + 8
End Function

Public Function GUID_WICPixelFormat48bppRGB() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat48bppRGB
    GetMem8 156994148506150.6481@, ByVal VarPtr(GUID_WICPixelFormat48bppRGB) + 8
End Function

Public Function GUID_WICPixelFormat48bppBGR() As UUID
    GetMem8 510221379048607.834@, GUID_WICPixelFormat48bppBGR
    GetMem8 138820655163730.7067@, ByVal VarPtr(GUID_WICPixelFormat48bppBGR) + 8
End Function

Public Function GUID_WICPixelFormat64bppRGBA() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat64bppRGBA
    GetMem8 164199907909943.4417@, ByVal VarPtr(GUID_WICPixelFormat64bppRGBA) + 8
End Function

Public Function GUID_WICPixelFormat64bppBGRA() As UUID
    GetMem8 511435120135549.734@, GUID_WICPixelFormat64bppBGRA
    GetMem8 505373523486930.4983@, ByVal VarPtr(GUID_WICPixelFormat64bppBGRA) + 8
End Function

Public Function GUID_WICPixelFormat64bppPRGBA() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat64bppPRGBA
    GetMem8 171405667313736.2353@, ByVal VarPtr(GUID_WICPixelFormat64bppPRGBA) + 8
End Function

Public Function GUID_WICPixelFormat64bppPBGRA() As UUID
    GetMem8 508333794029112.8974@, GUID_WICPixelFormat64bppPBGRA
    GetMem8 348286179994982.4174@, ByVal VarPtr(GUID_WICPixelFormat64bppPBGRA) + 8
End Function

Public Function GUID_WICPixelFormat16bppGrayFixedPoint() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat16bppGrayFixedPoint
    GetMem8 142582629698565.0609@, ByVal VarPtr(GUID_WICPixelFormat16bppGrayFixedPoint) + 8
End Function

Public Function GUID_WICPixelFormat32bppBGR101010() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat32bppBGR101010
    GetMem8 149788389102357.8545@, ByVal VarPtr(GUID_WICPixelFormat32bppBGR101010) + 8
End Function

Public Function GUID_WICPixelFormat48bppRGBFixedPoint() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat48bppRGBFixedPoint
    GetMem8 135376870294772.2673@, ByVal VarPtr(GUID_WICPixelFormat48bppRGBFixedPoint) + 8
End Function

Public Function GUID_WICPixelFormat48bppBGRFixedPoint() As UUID
    GetMem8 527703427266550.683@, GUID_WICPixelFormat48bppBGRFixedPoint
    GetMem8 304984237878443.2029@, ByVal VarPtr(GUID_WICPixelFormat48bppBGRFixedPoint) + 8
End Function

Public Function GUID_WICPixelFormat96bppRGBFixedPoint() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat96bppRGBFixedPoint
    GetMem8 178611426717529.0289@, ByVal VarPtr(GUID_WICPixelFormat96bppRGBFixedPoint) + 8
End Function

Public Function GUID_WICPixelFormat128bppRGBAFloat() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat128bppRGBAFloat
    GetMem8 185817186121321.8225@, ByVal VarPtr(GUID_WICPixelFormat128bppRGBAFloat) + 8
End Function

Public Function GUID_WICPixelFormat128bppPRGBAFloat() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat128bppPRGBAFloat
    GetMem8 193022945525114.6161@, ByVal VarPtr(GUID_WICPixelFormat128bppPRGBAFloat) + 8
End Function

Public Function GUID_WICPixelFormat128bppRGBFloat() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat128bppRGBFloat
    GetMem8 200228704928907.4097@, ByVal VarPtr(GUID_WICPixelFormat128bppRGBFloat) + 8
End Function

Public Function GUID_WICPixelFormat32bppCMYK() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat32bppCMYK
    GetMem8 207434464332700.2033@, ByVal VarPtr(GUID_WICPixelFormat32bppCMYK) + 8
End Function

Public Function GUID_WICPixelFormat64bppRGBAFixedPoint() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat64bppRGBAFixedPoint
    GetMem8 214640223736492.9969@, ByVal VarPtr(GUID_WICPixelFormat64bppRGBAFixedPoint) + 8
End Function

Public Function GUID_WICPixelFormat64bppBGRAFixedPoint() As UUID
    GetMem8 534220684480779.9612@, GUID_WICPixelFormat64bppBGRAFixedPoint
    GetMem8 330246011184814.6107@, ByVal VarPtr(GUID_WICPixelFormat64bppBGRAFixedPoint) + 8
End Function

Public Function GUID_WICPixelFormat64bppRGBFixedPoint() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat64bppRGBFixedPoint
    GetMem8 466841802869240.7729@, ByVal VarPtr(GUID_WICPixelFormat64bppRGBFixedPoint) + 8
End Function

Public Function GUID_WICPixelFormat128bppRGBAFixedPoint() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat128bppRGBAFixedPoint
    GetMem8 221845983140285.7905@, ByVal VarPtr(GUID_WICPixelFormat128bppRGBAFixedPoint) + 8
End Function

Public Function GUID_WICPixelFormat128bppRGBFixedPoint() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat128bppRGBFixedPoint
    GetMem8 474047562273033.5665@, ByVal VarPtr(GUID_WICPixelFormat128bppRGBFixedPoint) + 8
End Function

Public Function GUID_WICPixelFormat64bppRGBAHalf() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat64bppRGBAHalf
    GetMem8 423607246446484.0113@, ByVal VarPtr(GUID_WICPixelFormat64bppRGBAHalf) + 8
End Function

Public Function GUID_WICPixelFormat64bppRGBHalf() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat64bppRGBHalf
    GetMem8 481253321676826.3601@, ByVal VarPtr(GUID_WICPixelFormat64bppRGBHalf) + 8
End Function

Public Function GUID_WICPixelFormat48bppRGBHalf() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat48bppRGBHalf
    GetMem8 430813005850276.8049@, ByVal VarPtr(GUID_WICPixelFormat48bppRGBHalf) + 8
End Function

Public Function GUID_WICPixelFormat32bppRGBE() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat32bppRGBE
    GetMem8 445224524657862.3921@, ByVal VarPtr(GUID_WICPixelFormat32bppRGBE) + 8
End Function

Public Function GUID_WICPixelFormat16bppGrayHalf() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat16bppGrayHalf
    GetMem8 452430284061655.1857@, ByVal VarPtr(GUID_WICPixelFormat16bppGrayHalf) + 8
End Function

Public Function GUID_WICPixelFormat32bppGrayFixedPoint() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat32bppGrayFixedPoint
    GetMem8 459636043465447.9793@, ByVal VarPtr(GUID_WICPixelFormat32bppGrayFixedPoint) + 8
End Function

Public Function GUID_WICPixelFormat32bppRGBA1010102() As UUID
    GetMem8 498182228482533.3106@, GUID_WICPixelFormat32bppRGBA1010102
    GetMem8 -228172643511533.0379@, ByVal VarPtr(GUID_WICPixelFormat32bppRGBA1010102) + 8
End Function

Public Function GUID_WICPixelFormat32bppRGBA1010102XR() As UUID
    GetMem8 484918163384817.5514@, GUID_WICPixelFormat32bppRGBA1010102XR
    GetMem8 317584848147552.7349@, ByVal VarPtr(GUID_WICPixelFormat32bppRGBA1010102XR) + 8
End Function

Public Function GUID_WICPixelFormat64bppCMYK() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat64bppCMYK
    GetMem8 229051742544078.5841@, ByVal VarPtr(GUID_WICPixelFormat64bppCMYK) + 8
End Function

Public Function GUID_WICPixelFormat24bpp3Channels() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat24bpp3Channels
    GetMem8 236257501947871.3777@, ByVal VarPtr(GUID_WICPixelFormat24bpp3Channels) + 8
End Function

Public Function GUID_WICPixelFormat32bpp4Channels() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat32bpp4Channels
    GetMem8 243463261351664.1713@, ByVal VarPtr(GUID_WICPixelFormat32bpp4Channels) + 8
End Function

Public Function GUID_WICPixelFormat40bpp5Channels() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat40bpp5Channels
    GetMem8 250669020755456.9649@, ByVal VarPtr(GUID_WICPixelFormat40bpp5Channels) + 8
End Function

Public Function GUID_WICPixelFormat48bpp6Channels() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat48bpp6Channels
    GetMem8 257874780159249.7585@, ByVal VarPtr(GUID_WICPixelFormat48bpp6Channels) + 8
End Function

Public Function GUID_WICPixelFormat56bpp7Channels() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat56bpp7Channels
    GetMem8 265080539563042.5521@, ByVal VarPtr(GUID_WICPixelFormat56bpp7Channels) + 8
End Function

Public Function GUID_WICPixelFormat64bpp8Channels() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat64bpp8Channels
    GetMem8 272286298966835.3457@, ByVal VarPtr(GUID_WICPixelFormat64bpp8Channels) + 8
End Function

Public Function GUID_WICPixelFormat48bpp3Channels() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat48bpp3Channels
    GetMem8 279492058370628.1393@, ByVal VarPtr(GUID_WICPixelFormat48bpp3Channels) + 8
End Function

Public Function GUID_WICPixelFormat64bpp4Channels() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat64bpp4Channels
    GetMem8 286697817774420.9329@, ByVal VarPtr(GUID_WICPixelFormat64bpp4Channels) + 8
End Function

Public Function GUID_WICPixelFormat80bpp5Channels() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat80bpp5Channels
    GetMem8 293903577178213.7265@, ByVal VarPtr(GUID_WICPixelFormat80bpp5Channels) + 8
End Function

Public Function GUID_WICPixelFormat96bpp6Channels() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat96bpp6Channels
    GetMem8 301109336582006.5201@, ByVal VarPtr(GUID_WICPixelFormat96bpp6Channels) + 8
End Function

Public Function GUID_WICPixelFormat112bpp7Channels() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat112bpp7Channels
    GetMem8 308315095985799.3137@, ByVal VarPtr(GUID_WICPixelFormat112bpp7Channels) + 8
End Function

Public Function GUID_WICPixelFormat128bpp8Channels() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat128bpp8Channels
    GetMem8 315520855389592.1073@, ByVal VarPtr(GUID_WICPixelFormat128bpp8Channels) + 8
End Function

Public Function GUID_WICPixelFormat40bppCMYKAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat40bppCMYKAlpha
    GetMem8 322726614793384.9009@, ByVal VarPtr(GUID_WICPixelFormat40bppCMYKAlpha) + 8
End Function

Public Function GUID_WICPixelFormat80bppCMYKAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat80bppCMYKAlpha
    GetMem8 329932374197177.6945@, ByVal VarPtr(GUID_WICPixelFormat80bppCMYKAlpha) + 8
End Function

Public Function GUID_WICPixelFormat32bpp3ChannelsAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat32bpp3ChannelsAlpha
    GetMem8 337138133600970.4881@, ByVal VarPtr(GUID_WICPixelFormat32bpp3ChannelsAlpha) + 8
End Function

Public Function GUID_WICPixelFormat40bpp4ChannelsAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat40bpp4ChannelsAlpha
    GetMem8 344343893004763.2817@, ByVal VarPtr(GUID_WICPixelFormat40bpp4ChannelsAlpha) + 8
End Function

Public Function GUID_WICPixelFormat48bpp5ChannelsAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat48bpp5ChannelsAlpha
    GetMem8 351549652408556.0753@, ByVal VarPtr(GUID_WICPixelFormat48bpp5ChannelsAlpha) + 8
End Function

Public Function GUID_WICPixelFormat56bpp6ChannelsAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat56bpp6ChannelsAlpha
    GetMem8 358755411812348.8689@, ByVal VarPtr(GUID_WICPixelFormat56bpp6ChannelsAlpha) + 8
End Function

Public Function GUID_WICPixelFormat64bpp7ChannelsAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat64bpp7ChannelsAlpha
    GetMem8 365961171216141.6625@, ByVal VarPtr(GUID_WICPixelFormat64bpp7ChannelsAlpha) + 8
End Function

Public Function GUID_WICPixelFormat72bpp8ChannelsAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat72bpp8ChannelsAlpha
    GetMem8 373166930619934.4561@, ByVal VarPtr(GUID_WICPixelFormat72bpp8ChannelsAlpha) + 8
End Function

Public Function GUID_WICPixelFormat64bpp3ChannelsAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat64bpp3ChannelsAlpha
    GetMem8 380372690023727.2497@, ByVal VarPtr(GUID_WICPixelFormat64bpp3ChannelsAlpha) + 8
End Function

Public Function GUID_WICPixelFormat80bpp4ChannelsAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat80bpp4ChannelsAlpha
    GetMem8 387578449427520.0433@, ByVal VarPtr(GUID_WICPixelFormat80bpp4ChannelsAlpha) + 8
End Function

Public Function GUID_WICPixelFormat96bpp5ChannelsAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat96bpp5ChannelsAlpha
    GetMem8 394784208831312.8369@, ByVal VarPtr(GUID_WICPixelFormat96bpp5ChannelsAlpha) + 8
End Function

Public Function GUID_WICPixelFormat112bpp6ChannelsAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat112bpp6ChannelsAlpha
    GetMem8 401989968235105.6305@, ByVal VarPtr(GUID_WICPixelFormat112bpp6ChannelsAlpha) + 8
End Function

Public Function GUID_WICPixelFormat128bpp7ChannelsAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat128bpp7ChannelsAlpha
    GetMem8 409195727638898.4241@, ByVal VarPtr(GUID_WICPixelFormat128bpp7ChannelsAlpha) + 8
End Function

Public Function GUID_WICPixelFormat144bpp8ChannelsAlpha() As UUID
    GetMem8 547589997359777.4628@, GUID_WICPixelFormat144bpp8ChannelsAlpha
    GetMem8 416401487042691.2177@, ByVal VarPtr(GUID_WICPixelFormat144bpp8ChannelsAlpha) + 8
End Function

Public Function GUID_MetadataFormatUnknown() As UUID
    GetMem8 536732370374691.0511@, GUID_MetadataFormatUnknown
    GetMem8 224162087803386.0013@, ByVal VarPtr(GUID_MetadataFormatUnknown) + 8
End Function

Public Function GUID_MetadataFormatIfd() As UUID
    GetMem8 545559807073475.5526@, GUID_MetadataFormatIfd
    GetMem8 -236152826505776.5221@, ByVal VarPtr(GUID_MetadataFormatIfd) + 8
End Function

Public Function GUID_MetadataFormatSubIfd() As UUID
    GetMem8 564503093201147.524@, GUID_MetadataFormatSubIfd
    GetMem8 359024690370557.2539@, ByVal VarPtr(GUID_MetadataFormatSubIfd) + 8
End Function

Public Function GUID_MetadataFormatExif() As UUID
    GetMem8 507941858318459.2797@, GUID_MetadataFormatExif
    GetMem8 633497449761017.538@, ByVal VarPtr(GUID_MetadataFormatExif) + 8
End Function

Public Function GUID_MetadataFormatGps() As UUID
    GetMem8 494877354355095.4378@, GUID_MetadataFormatGps
    GetMem8 -144038913514639.2913@, ByVal VarPtr(GUID_MetadataFormatGps) + 8
End Function

Public Function GUID_MetadataFormatInterop() As UUID
    GetMem8 551561665498162.3694@, GUID_MetadataFormatInterop
    GetMem8 -23757493194874.8355@, ByVal VarPtr(GUID_MetadataFormatInterop) + 8
End Function

Public Function GUID_MetadataFormatApp0() As UUID
    GetMem8 503225202269542.404@, GUID_MetadataFormatApp0
    GetMem8 -394196863132383.5741@, ByVal VarPtr(GUID_MetadataFormatApp0) + 8
End Function

Public Function GUID_MetadataFormatApp1() As UUID
    GetMem8 527258191746799.2003@, GUID_MetadataFormatApp1
    GetMem8 -571792456674957.7343@, ByVal VarPtr(GUID_MetadataFormatApp1) + 8
End Function

Public Function GUID_MetadataFormatApp13() As UUID
    GetMem8 485177208836910.8642@, GUID_MetadataFormatApp13
    GetMem8 -533481911254890.89@, ByVal VarPtr(GUID_MetadataFormatApp13) + 8
End Function

Public Function GUID_MetadataFormatIPTC() As UUID
    GetMem8 464993270782984.4244@, GUID_MetadataFormatIPTC
    GetMem8 -535723716998143.5487@, ByVal VarPtr(GUID_MetadataFormatIPTC) + 8
End Function

Public Function GUID_MetadataFormatIRB() As UUID
    GetMem8 545653913998904.8678@, GUID_MetadataFormatIRB
    GetMem8 747998496810212.9081@, ByVal VarPtr(GUID_MetadataFormatIRB) + 8
End Function

Public Function GUID_MetadataFormat8BIMIPTC() As UUID
    GetMem8 565033783077111.9756@, GUID_MetadataFormat8BIMIPTC
    GetMem8 345999120888159.0705@, ByVal VarPtr(GUID_MetadataFormat8BIMIPTC) + 8
End Function

Public Function GUID_MetadataFormat8BIMResolutionInfo() As UUID
    GetMem8 488514100035106.8253@, GUID_MetadataFormat8BIMResolutionInfo
    GetMem8 28400082123000.1836@, ByVal VarPtr(GUID_MetadataFormat8BIMResolutionInfo) + 8
End Function

Public Function GUID_MetadataFormat8BIMIPTCDigest() As UUID
    GetMem8 515397922833479.3349@, GUID_MetadataFormat8BIMIPTCDigest
    GetMem8 47758234732788.7499@, ByVal VarPtr(GUID_MetadataFormat8BIMIPTCDigest) + 8
End Function

Public Function GUID_MetadataFormatXMP() As UUID
    GetMem8 554307142083273.0168@, GUID_MetadataFormatXMP
    GetMem8 -624097813657657.1994@, ByVal VarPtr(GUID_MetadataFormatXMP) + 8
End Function

Public Function GUID_MetadataFormatThumbnail() As UUID
    GetMem8 467882551044720.6121@, GUID_MetadataFormatThumbnail
    GetMem8 -835707122058926.885@, ByVal VarPtr(GUID_MetadataFormatThumbnail) + 8
End Function

Public Function GUID_MetadataFormatChunktEXt() As UUID
    GetMem8 527026782248773.4582@, GUID_MetadataFormatChunktEXt
    GetMem8 -485962674874827.6336@, ByVal VarPtr(GUID_MetadataFormatChunktEXt) + 8
End Function

Public Function GUID_MetadataFormatXMPStruct() As UUID
    GetMem8 563370086750121.0865@, GUID_MetadataFormatXMPStruct
    GetMem8 -344513545147304.5585@, ByVal VarPtr(GUID_MetadataFormatXMPStruct) + 8
End Function

Public Function GUID_MetadataFormatXMPBag() As UUID
    GetMem8 497840911884358.7167@, GUID_MetadataFormatXMPBag
    GetMem8 -195564561978114.4704@, ByVal VarPtr(GUID_MetadataFormatXMPBag) + 8
End Function

Public Function GUID_MetadataFormatXMPSeq() As UUID
    GetMem8 500263213686697.9586@, GUID_MetadataFormatXMPSeq
    GetMem8 524846979844526.1986@, ByVal VarPtr(GUID_MetadataFormatXMPSeq) + 8
End Function

Public Function GUID_MetadataFormatXMPAlt() As UUID
    GetMem8 519590675649663.1413@, GUID_MetadataFormatXMPAlt
    GetMem8 427871023344475.1527@, ByVal VarPtr(GUID_MetadataFormatXMPAlt) + 8
End Function

Public Function GUID_MetadataFormatLSD() As UUID
    GetMem8 527185325188068.227@, GUID_MetadataFormatLSD
    GetMem8 -787378801367196.9351@, ByVal VarPtr(GUID_MetadataFormatLSD) + 8
End Function

Public Function GUID_MetadataFormatIMD() As UUID
    GetMem8 525043775834028.0454@, GUID_MetadataFormatIMD
    GetMem8 -809338497778139.3514@, ByVal VarPtr(GUID_MetadataFormatIMD) + 8
End Function

Public Function GUID_MetadataFormatGCE() As UUID
    GetMem8 550617712104294.268@, GUID_MetadataFormatGCE
    GetMem8 -15919482358631.4073@, ByVal VarPtr(GUID_MetadataFormatGCE) + 8
End Function

Public Function GUID_MetadataFormatAPE() As UUID
    GetMem8 562212115483277.6642@, GUID_MetadataFormatAPE
    GetMem8 -435793746704561.8041@, ByVal VarPtr(GUID_MetadataFormatAPE) + 8
End Function

Public Function GUID_MetadataFormatJpegChrominance() As UUID
    GetMem8 573021345484564.4239@, GUID_MetadataFormatJpegChrominance
    GetMem8 -59484311688288.2917@, ByVal VarPtr(GUID_MetadataFormatJpegChrominance) + 8
End Function

Public Function GUID_MetadataFormatJpegLuminance() As UUID
    GetMem8 521543003734019.2775@, GUID_MetadataFormatJpegLuminance
    GetMem8 636815904311910.6957@, ByVal VarPtr(GUID_MetadataFormatJpegLuminance) + 8
End Function

Public Function GUID_MetadataFormatJpegComment() As UUID
    GetMem8 513823754622064.2099@, GUID_MetadataFormatJpegComment
    GetMem8 633802582062302.4541@, ByVal VarPtr(GUID_MetadataFormatJpegComment) + 8
End Function

Public Function GUID_MetadataFormatGifComment() As UUID
    GetMem8 539188155419399.8048@, GUID_MetadataFormatGifComment
    GetMem8 537823599173528.4651@, ByVal VarPtr(GUID_MetadataFormatGifComment) + 8
End Function

Public Function GUID_MetadataFormatChunkgAMA() As UUID
    GetMem8 553523770631134.9669@, GUID_MetadataFormatChunkgAMA
    GetMem8 -910404771293038.5279@, ByVal VarPtr(GUID_MetadataFormatChunkgAMA) + 8
End Function

Public Function GUID_MetadataFormatChunkbKGD() As UUID
    GetMem8 561441784193752.8177@, GUID_MetadataFormatChunkbKGD
    GetMem8 -519730345715230.0362@, ByVal VarPtr(GUID_MetadataFormatChunkbKGD) + 8
End Function

Public Function GUID_MetadataFormatChunkiTXt() As UUID
    GetMem8 543782761964495.0313@, GUID_MetadataFormatChunkiTXt
    GetMem8 144809711167039.4538@, ByVal VarPtr(GUID_MetadataFormatChunkiTXt) + 8
End Function

Public Function GUID_MetadataFormatChunkcHRM() As UUID
    GetMem8 495034468198903.9451@, GUID_MetadataFormatChunkcHRM
    GetMem8 766215975661066.0224@, ByVal VarPtr(GUID_MetadataFormatChunkcHRM) + 8
End Function

Public Function GUID_MetadataFormatChunkhIST() As UUID
    GetMem8 523454996148928.585@, GUID_MetadataFormatChunkhIST
    GetMem8 -764283585039286.8163@, ByVal VarPtr(GUID_MetadataFormatChunkhIST) + 8
End Function

Public Function GUID_MetadataFormatChunkiCCP() As UUID
    GetMem8 497639679956164.8555@, GUID_MetadataFormatChunkiCCP
    GetMem8 780574410434207.4769@, ByVal VarPtr(GUID_MetadataFormatChunkiCCP) + 8
End Function

Public Function GUID_MetadataFormatChunksRGB() As UUID
    GetMem8 563845003884403.0262@, GUID_MetadataFormatChunksRGB
    GetMem8 -276048828719153.8813@, ByVal VarPtr(GUID_MetadataFormatChunksRGB) + 8
End Function

Public Function GUID_MetadataFormatChunktIME() As UUID
    GetMem8 504709514596769.3357@, GUID_MetadataFormatChunktIME
    GetMem8 -18403096251651.108@, ByVal VarPtr(GUID_MetadataFormatChunktIME) + 8
End Function

Public Function CLSID_WICUnknownMetadataReader() As UUID
    GetMem8 544099970060538.8226@, CLSID_WICUnknownMetadataReader
    GetMem8 -829201150537230.652@, ByVal VarPtr(CLSID_WICUnknownMetadataReader) + 8
End Function

Public Function CLSID_WICUnknownMetadataWriter() As UUID
    GetMem8 570863768518084.4678@, CLSID_WICUnknownMetadataWriter
    GetMem8 -28573597888398.4496@, ByVal VarPtr(CLSID_WICUnknownMetadataWriter) + 8
End Function

Public Function CLSID_WICApp0MetadataWriter() As UUID
    GetMem8 530025164035890.8834@, CLSID_WICApp0MetadataWriter
    GetMem8 -239701072335203.6465@, ByVal VarPtr(CLSID_WICApp0MetadataWriter) + 8
End Function

Public Function CLSID_WICApp0MetadataReader() As UUID
    GetMem8 519255312913100.2675@, CLSID_WICApp0MetadataReader
    GetMem8 365939972870439.3617@, ByVal VarPtr(CLSID_WICApp0MetadataReader) + 8
End Function

Public Function CLSID_WICApp1MetadataWriter() As UUID
    GetMem8 476004993817787.6073@, CLSID_WICApp1MetadataWriter
    GetMem8 183269091549252.8563@, ByVal VarPtr(CLSID_WICApp1MetadataWriter) + 8
End Function

Public Function CLSID_WICApp1MetadataReader() As UUID
    GetMem8 546215310368408.9107@, CLSID_WICApp1MetadataReader
    GetMem8 -26036580535746.7218@, ByVal VarPtr(CLSID_WICApp1MetadataReader) + 8
End Function

Public Function CLSID_WICApp13MetadataWriter() As UUID
    GetMem8 532484887308885.4297@, CLSID_WICApp13MetadataWriter
    GetMem8 -307699834521476.9731@, ByVal VarPtr(CLSID_WICApp13MetadataWriter) + 8
End Function

Public Function CLSID_WICApp13MetadataReader() As UUID
    GetMem8 504530514639783.432@, CLSID_WICApp13MetadataReader
    GetMem8 -70484142198600.378@, ByVal VarPtr(CLSID_WICApp13MetadataReader) + 8
End Function

Public Function CLSID_WICIfdMetadataReader() As UUID
    GetMem8 567076754949677.2182@, CLSID_WICIfdMetadataReader
    GetMem8 -182887219373152.0112@, ByVal VarPtr(CLSID_WICIfdMetadataReader) + 8
End Function

Public Function CLSID_WICIfdMetadataWriter() As UUID
    GetMem8 516190993949104.4392@, CLSID_WICIfdMetadataWriter
    GetMem8 -637946381286498.6227@, ByVal VarPtr(CLSID_WICIfdMetadataWriter) + 8
End Function

Public Function CLSID_WICSubIfdMetadataReader() As UUID
    GetMem8 542287581007919.0793@, CLSID_WICSubIfdMetadataReader
    GetMem8 715809321303109.5734@, ByVal VarPtr(CLSID_WICSubIfdMetadataReader) + 8
End Function

Public Function CLSID_WICSubIfdMetadataWriter() As UUID
    GetMem8 571409882592721.8054@, CLSID_WICSubIfdMetadataWriter
    GetMem8 408533498888460.3564@, ByVal VarPtr(CLSID_WICSubIfdMetadataWriter) + 8
End Function

Public Function CLSID_WICExifMetadataReader() As UUID
    GetMem8 535285526118899.5168@, CLSID_WICExifMetadataReader
    GetMem8 480205161997111.1871@, ByVal VarPtr(CLSID_WICExifMetadataReader) + 8
End Function

Public Function CLSID_WICExifMetadataWriter() As UUID
    GetMem8 504734246036211.6314@, CLSID_WICExifMetadataWriter
    GetMem8 -794463700353571.416@, ByVal VarPtr(CLSID_WICExifMetadataWriter) + 8
End Function

Public Function CLSID_WICGpsMetadataReader() As UUID
    GetMem8 521013945662855.3995@, CLSID_WICGpsMetadataReader
    GetMem8 885888895961976.9753@, ByVal VarPtr(CLSID_WICGpsMetadataReader) + 8
End Function

Public Function CLSID_WICGpsMetadataWriter() As UUID
    GetMem8 551870692633267.914@, CLSID_WICGpsMetadataWriter
    GetMem8 854695652350830.4804@, ByVal VarPtr(CLSID_WICGpsMetadataWriter) + 8
End Function

Public Function CLSID_WICInteropMetadataReader() As UUID
    GetMem8 501672901118005.8776@, CLSID_WICInteropMetadataReader
    GetMem8 150710638662529.8615@, ByVal VarPtr(CLSID_WICInteropMetadataReader) + 8
End Function

Public Function CLSID_WICInteropMetadataWriter() As UUID
    GetMem8 496094093090323.0021@, CLSID_WICInteropMetadataWriter
    GetMem8 113202542591164.5873@, ByVal VarPtr(CLSID_WICInteropMetadataWriter) + 8
End Function

Public Function CLSID_WICThumbnailMetadataReader() As UUID
    GetMem8 496070285618218.0185@, CLSID_WICThumbnailMetadataReader
    GetMem8 633084010118368.7069@, ByVal VarPtr(CLSID_WICThumbnailMetadataReader) + 8
End Function

Public Function CLSID_WICThumbnailMetadataWriter() As UUID
    GetMem8 497151419009267.7644@, CLSID_WICThumbnailMetadataWriter
    GetMem8 -916457149257222.0496@, ByVal VarPtr(CLSID_WICThumbnailMetadataWriter) + 8
End Function

Public Function CLSID_WICIPTCMetadataReader() As UUID
    GetMem8 496070285202143.0617@, CLSID_WICIPTCMetadataReader
    GetMem8 633084010118368.7069@, ByVal VarPtr(CLSID_WICIPTCMetadataReader) + 8
End Function

Public Function CLSID_WICIPTCMetadataWriter() As UUID
    GetMem8 497151418690500.6604@, CLSID_WICIPTCMetadataWriter
    GetMem8 -916457149257222.0496@, ByVal VarPtr(CLSID_WICIPTCMetadataWriter) + 8
End Function

Public Function CLSID_WICIRBMetadataReader() As UUID
    GetMem8 517736799552699.2855@, CLSID_WICIRBMetadataReader
    GetMem8 -665502871991033.865@, ByVal VarPtr(CLSID_WICIRBMetadataReader) + 8
End Function

Public Function CLSID_WICIRBMetadataWriter() As UUID
    GetMem8 491455552157411.5637@, CLSID_WICIRBMetadataWriter
    GetMem8 -416303606639901.1712@, ByVal VarPtr(CLSID_WICIRBMetadataWriter) + 8
End Function

Public Function CLSID_WIC8BIMIPTCMetadataReader() As UUID
    GetMem8 559516838744348.43@, CLSID_WIC8BIMIPTCMetadataReader
    GetMem8 -808319812232841.5068@, ByVal VarPtr(CLSID_WIC8BIMIPTCMetadataReader) + 8
End Function

Public Function CLSID_WIC8BIMIPTCMetadataWriter() As UUID
    GetMem8 494577730374759.2742@, CLSID_WIC8BIMIPTCMetadataWriter
    GetMem8 -361563201860877.6034@, ByVal VarPtr(CLSID_WIC8BIMIPTCMetadataWriter) + 8
End Function

Public Function CLSID_WIC8BIMResolutionInfoMetadataReader() As UUID
    GetMem8 572770272596230.6426@, CLSID_WIC8BIMResolutionInfoMetadataReader
    GetMem8 -742042520762044.9101@, ByVal VarPtr(CLSID_WIC8BIMResolutionInfoMetadataReader) + 8
End Function

Public Function CLSID_WIC8BIMResolutionInfoMetadataWriter() As UUID
    GetMem8 543638053156782.0302@, CLSID_WIC8BIMResolutionInfoMetadataWriter
    GetMem8 -504194717708322.2888@, ByVal VarPtr(CLSID_WIC8BIMResolutionInfoMetadataWriter) + 8
End Function

Public Function CLSID_WIC8BIMIPTCDigestMetadataReader() As UUID
    GetMem8 470959276150910.9534@, CLSID_WIC8BIMIPTCDigestMetadataReader
    GetMem8 -644671682684480.5758@, ByVal VarPtr(CLSID_WIC8BIMIPTCDigestMetadataReader) + 8
End Function

Public Function CLSID_WIC8BIMIPTCDigestMetadataWriter() As UUID
    GetMem8 528695922435593.9883@, CLSID_WIC8BIMIPTCDigestMetadataWriter
    GetMem8 -603270573416286.6801@, ByVal VarPtr(CLSID_WIC8BIMIPTCDigestMetadataWriter) + 8
End Function

Public Function CLSID_WICPngTextMetadataReader() As UUID
    GetMem8 465073271413576.0844@, CLSID_WICPngTextMetadataReader
    GetMem8 -634171151155872.1354@, ByVal VarPtr(CLSID_WICPngTextMetadataReader) + 8
End Function

Public Function CLSID_WICXMPMetadataWriter() As UUID
    GetMem8 505701002932321.5182@, CLSID_WICXMPMetadataWriter
    GetMem8 -416747568845857.1338@, ByVal VarPtr(CLSID_WICXMPMetadataWriter) + 8
End Function

Public Function CLSID_WICXMPStructMetadataReader() As UUID
    GetMem8 518575647113563.689@, CLSID_WICXMPStructMetadataReader
    GetMem8 -136544688391000.4068@, ByVal VarPtr(CLSID_WICXMPStructMetadataReader) + 8
End Function

Public Function CLSID_WICXMPStructMetadataWriter() As UUID
    GetMem8 469176329194767.1443@, CLSID_WICXMPStructMetadataWriter
    GetMem8 -487907983440077.8341@, ByVal VarPtr(CLSID_WICXMPStructMetadataWriter) + 8
End Function

Public Function CLSID_WICXMPBagMetadataReader() As UUID
    GetMem8 574076920430169.9632@, CLSID_WICXMPBagMetadataReader
    GetMem8 -470220311858138.3027@, ByVal VarPtr(CLSID_WICXMPBagMetadataReader) + 8
End Function

Public Function CLSID_WICXMPBagMetadataWriter() As UUID
    GetMem8 482837639103474.3948@, CLSID_WICXMPBagMetadataWriter
    GetMem8 -808319377718809.7626@, ByVal VarPtr(CLSID_WICXMPBagMetadataWriter) + 8
End Function

Public Function CLSID_WICXMPSeqMetadataReader() As UUID
    GetMem8 488865348492741.2051@, CLSID_WICXMPSeqMetadataReader
    GetMem8 -535605610360483.9003@, ByVal VarPtr(CLSID_WICXMPSeqMetadataReader) + 8
End Function

Public Function CLSID_WICXMPSeqMetadataWriter() As UUID
    GetMem8 540877499054429.4366@, CLSID_WICXMPSeqMetadataWriter
    GetMem8 -635974071838564.491@, ByVal VarPtr(CLSID_WICXMPSeqMetadataWriter) + 8
End Function

Public Function CLSID_WICXMPAltMetadataReader() As UUID
    GetMem8 523113403610646.8546@, CLSID_WICXMPAltMetadataReader
    GetMem8 -783518179708736.5704@, ByVal VarPtr(CLSID_WICXMPAltMetadataReader) + 8
End Function

Public Function CLSID_WICXMPAltMetadataWriter() As UUID
    GetMem8 549635258892918.2316@, CLSID_WICXMPAltMetadataWriter
    GetMem8 -155204323115079.1769@, ByVal VarPtr(CLSID_WICXMPAltMetadataWriter) + 8
End Function

Public Function CLSID_WICLSDMetadataReader() As UUID
    GetMem8 515953515998471.9763@, CLSID_WICLSDMetadataReader
    GetMem8 -21921748317857.5967@, ByVal VarPtr(CLSID_WICLSDMetadataReader) + 8
End Function

Public Function CLSID_WICLSDMetadataWriter() As UUID
    GetMem8 528410098492507.5431@, CLSID_WICLSDMetadataWriter
    GetMem8 751859917657718.2343@, ByVal VarPtr(CLSID_WICLSDMetadataWriter) + 8
End Function

Public Function CLSID_WICGCEMetadataReader() As UUID
    GetMem8 475241160853515.1709@, CLSID_WICGCEMetadataReader
    GetMem8 -505206815750197.1787@, ByVal VarPtr(CLSID_WICGCEMetadataReader) + 8
End Function

Public Function CLSID_WICGCEMetadataWriter() As UUID
    GetMem8 518479402771610.7382@, CLSID_WICGCEMetadataWriter
    GetMem8 -175995035866858.0173@, ByVal VarPtr(CLSID_WICGCEMetadataWriter) + 8
End Function

Public Function CLSID_WICIMDMetadataReader() As UUID
    GetMem8 481209629399053.9879@, CLSID_WICIMDMetadataReader
    GetMem8 702667818359193.2328@, ByVal VarPtr(CLSID_WICIMDMetadataReader) + 8
End Function

Public Function CLSID_WICIMDMetadataWriter() As UUID
    GetMem8 566250817271686.9407@, CLSID_WICIMDMetadataWriter
    GetMem8 824648029972919.1574@, ByVal VarPtr(CLSID_WICIMDMetadataWriter) + 8
End Function

Public Function CLSID_WICAPEMetadataReader() As UUID
    GetMem8 496597519530247.609@, CLSID_WICAPEMetadataReader
    GetMem8 756378245405907.7522@, ByVal VarPtr(CLSID_WICAPEMetadataReader) + 8
End Function

Public Function CLSID_WICAPEMetadataWriter() As UUID
    GetMem8 520142069675477.3962@, CLSID_WICAPEMetadataWriter
    GetMem8 -822818072605876.539@, ByVal VarPtr(CLSID_WICAPEMetadataWriter) + 8
End Function

Public Function CLSID_WICJpegChrominanceMetadataReader() As UUID
    GetMem8 500489178326352.2891@, CLSID_WICJpegChrominanceMetadataReader
    GetMem8 -162772041247044.6957@, ByVal VarPtr(CLSID_WICJpegChrominanceMetadataReader) + 8
End Function

Public Function CLSID_WICJpegChrominanceMetadataWriter() As UUID
    GetMem8 531999846674500.1712@, CLSID_WICJpegChrominanceMetadataWriter
    GetMem8 707414514123249.423@, ByVal VarPtr(CLSID_WICJpegChrominanceMetadataWriter) + 8
End Function

Public Function CLSID_WICJpegLuminanceMetadataReader() As UUID
    GetMem8 512735438718049.8824@, CLSID_WICJpegLuminanceMetadataReader
    GetMem8 409602914811613.3049@, ByVal VarPtr(CLSID_WICJpegLuminanceMetadataReader) + 8
End Function

Public Function CLSID_WICJpegLuminanceMetadataWriter() As UUID
    GetMem8 506867169885527.9292@, CLSID_WICJpegLuminanceMetadataWriter
    GetMem8 547506739876380.7385@, ByVal VarPtr(CLSID_WICJpegLuminanceMetadataWriter) + 8
End Function

Public Function CLSID_WICJpegCommentMetadataReader() As UUID
    GetMem8 549815711769337.5612@, CLSID_WICJpegCommentMetadataReader
    GetMem8 57379281401140.6507@, ByVal VarPtr(CLSID_WICJpegCommentMetadataReader) + 8
End Function

Public Function CLSID_WICJpegCommentMetadataWriter() As UUID
    GetMem8 568194810242840.2543@, CLSID_WICJpegCommentMetadataWriter
    GetMem8 -320205619373435.0207@, ByVal VarPtr(CLSID_WICJpegCommentMetadataWriter) + 8
End Function

Public Function CLSID_WICGifCommentMetadataReader() As UUID
    GetMem8 573460609498437.9707@, CLSID_WICGifCommentMetadataReader
    GetMem8 644048080639606.3363@, ByVal VarPtr(CLSID_WICGifCommentMetadataReader) + 8
End Function

Public Function CLSID_WICGifCommentMetadataWriter() As UUID
    GetMem8 472336636349511.0652@, CLSID_WICGifCommentMetadataWriter
    GetMem8 -678610354713666.6193@, ByVal VarPtr(CLSID_WICGifCommentMetadataWriter) + 8
End Function

Public Function CLSID_WICPngGamaMetadataReader() As UUID
    GetMem8 485062364854398.4185@, CLSID_WICPngGamaMetadataReader
    GetMem8 -308157837698615.7154@, ByVal VarPtr(CLSID_WICPngGamaMetadataReader) + 8
End Function

Public Function CLSID_WICPngGamaMetadataWriter() As UUID
    GetMem8 510634013349035.5475@, CLSID_WICPngGamaMetadataWriter
    GetMem8 576428380017656.2097@, ByVal VarPtr(CLSID_WICPngGamaMetadataWriter) + 8
End Function

Public Function CLSID_WICPngBkgdMetadataReader() As UUID
    GetMem8 535928785175469.3798@, CLSID_WICPngBkgdMetadataReader
    GetMem8 -267311623183479.0499@, ByVal VarPtr(CLSID_WICPngBkgdMetadataReader) + 8
End Function

Public Function CLSID_WICPngBkgdMetadataWriter() As UUID
    GetMem8 491826689321913.2157@, CLSID_WICPngBkgdMetadataWriter
    GetMem8 -804361994266860.4741@, ByVal VarPtr(CLSID_WICPngBkgdMetadataWriter) + 8
End Function

Public Function CLSID_WICPngItxtMetadataReader() As UUID
    GetMem8 537258118191092.1978@, CLSID_WICPngItxtMetadataReader
    GetMem8 258804474308391.5145@, ByVal VarPtr(CLSID_WICPngItxtMetadataReader) + 8
End Function

Public Function CLSID_WICPngItxtMetadataWriter() As UUID
    GetMem8 561849487105403.0617@, CLSID_WICPngItxtMetadataWriter
    GetMem8 -136783648461998.5512@, ByVal VarPtr(CLSID_WICPngItxtMetadataWriter) + 8
End Function

Public Function CLSID_WICPngChrmMetadataReader() As UUID
    GetMem8 462356787353638.4822@, CLSID_WICPngChrmMetadataReader
    GetMem8 710206867580418.9085@, ByVal VarPtr(CLSID_WICPngChrmMetadataReader) + 8
End Function

Public Function CLSID_WICPngChrmMetadataWriter() As UUID
    GetMem8 565746015306284.3371@, CLSID_WICPngChrmMetadataWriter
    GetMem8 -293142268816916.8964@, ByVal VarPtr(CLSID_WICPngChrmMetadataWriter) + 8
End Function

Public Function CLSID_WICPngHistMetadataReader() As UUID
    GetMem8 494090957047477.3431@, CLSID_WICPngHistMetadataReader
    GetMem8 237496712953532.9671@, ByVal VarPtr(CLSID_WICPngHistMetadataReader) + 8
End Function

Public Function CLSID_WICPngHistMetadataWriter() As UUID
    GetMem8 493099209159893.9977@, CLSID_WICPngHistMetadataWriter
    GetMem8 -2077217095233.5425@, ByVal VarPtr(CLSID_WICPngHistMetadataWriter) + 8
End Function

Public Function CLSID_WICPngIccpMetadataReader() As UUID
    GetMem8 505551385113263.4683@, CLSID_WICPngIccpMetadataReader
    GetMem8 -567714107879866.5564@, ByVal VarPtr(CLSID_WICPngIccpMetadataReader) + 8
End Function

Public Function CLSID_WICPngIccpMetadataWriter() As UUID
    GetMem8 553156042467567.9839@, CLSID_WICPngIccpMetadataWriter
    GetMem8 -241111256476153.2265@, ByVal VarPtr(CLSID_WICPngIccpMetadataWriter) + 8
End Function

Public Function CLSID_WICPngSrgbMetadataReader() As UUID
    GetMem8 528450411712375.758@, CLSID_WICPngSrgbMetadataReader
    GetMem8 740232737929171.6003@, ByVal VarPtr(CLSID_WICPngSrgbMetadataReader) + 8
End Function

Public Function CLSID_WICPngSrgbMetadataWriter() As UUID
    GetMem8 517900754698202.055@, CLSID_WICPngSrgbMetadataWriter
    GetMem8 -907573326898232.4577@, ByVal VarPtr(CLSID_WICPngSrgbMetadataWriter) + 8
End Function

Public Function CLSID_WICPngTimeMetadataReader() As UUID
    GetMem8 569647287416591.949@, CLSID_WICPngTimeMetadataReader
    GetMem8 -569249655263659.4043@, ByVal VarPtr(CLSID_WICPngTimeMetadataReader) + 8
End Function

Public Function CLSID_WICPngTimeMetadataWriter() As UUID
    GetMem8 558944832467602.7392@, CLSID_WICPngTimeMetadataWriter
    GetMem8 -182978765736772.8502@, ByVal VarPtr(CLSID_WICPngTimeMetadataWriter) + 8
End Function



