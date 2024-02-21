Attribute VB_Name = "D2D1Effects"
' //
' // Module D2D1Effects.bas - Helpes functions for Direct2D effects
' // By The trick 2023 (c)
' //

Option Explicit

Private Declare Function GetMem8 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
                         
' // {6AA97485-6354-4CFC-908C-E4A74F62C96C}
Public Function CLSID_D2D12DAffineTransform() As UUID
    GetMem8 554741805523150.1445@, CLSID_D2D12DAffineTransform
    GetMem8 783890472067377.064@, ByVal VarPtr(CLSID_D2D12DAffineTransform) + 8
End Function

' // {C2844D0B-3D86-46E7-85BA-526C9240F3FB}
Public Function CLSID_D2D13DPerspectiveTransform() As UUID
    GetMem8 510911995127347.9435@, CLSID_D2D13DPerspectiveTransform
    GetMem8 -29181855322218.6363@, ByVal VarPtr(CLSID_D2D13DPerspectiveTransform) + 8
End Function

' // {E8467B04-EC61-4B8A-B5DE-D4D73DEBEA5A}
Public Function CLSID_D2D13DTransform() As UUID
    GetMem8 544342300488358.17@, CLSID_D2D13DTransform
    GetMem8 655130725881039.2245@, ByVal VarPtr(CLSID_D2D13DTransform) + 8
End Function

' // {FC151437-049A-4784-A24A-F1C4DAF20987}
Public Function CLSID_D2D1ArithmeticComposite() As UUID
    GetMem8 515324893731940.4599@, CLSID_D2D1ArithmeticComposite
    GetMem8 -871616858237794.4414@, ByVal VarPtr(CLSID_D2D1ArithmeticComposite) + 8
End Function

' // {913E2BE4-FDCF-4FE2-A5F0-2454F14FF408}
Public Function CLSID_D2D1Atlas() As UUID
    GetMem8 575644234166974.3588@, CLSID_D2D1Atlas
    GetMem8 64522854453824.3237@, ByVal VarPtr(CLSID_D2D1Atlas) + 8
End Function

' // {5FB6C24D-C6DD-4231-9404-50F4D5C3252D}
Public Function CLSID_D2D1BitmapSource() As UUID
    GetMem8 476981213445795.0797@, CLSID_D2D1BitmapSource
    GetMem8 325322162953938.6516@, ByVal VarPtr(CLSID_D2D1BitmapSource) + 8
End Function

' // {81C5B77B-13F8-4CDD-AD20-C890547AC65D}
Public Function CLSID_D2D1Blend() As UUID
    GetMem8 553860507478561.3691@, CLSID_D2D1Blend
    GetMem8 675722279454088.4141@, ByVal VarPtr(CLSID_D2D1Blend) + 8
End Function

' // {2A2D49C0-4ACF-43C7-8C6A-7C4A27874D27}
Public Function CLSID_D2D1Border() As UUID
    GetMem8 488395457453288.9024@, CLSID_D2D1Border
    GetMem8 283206834350905.2044@, ByVal VarPtr(CLSID_D2D1Border) + 8
End Function

' // {8CEA8D1E-77B0-4986-B3B9-2F0C0EAE7887}
Public Function CLSID_D2D1Brightness() As UUID
    GetMem8 529805361181009.8462@, CLSID_D2D1Brightness
    GetMem8 -868500050602677.2045@, ByVal VarPtr(CLSID_D2D1Brightness) + 8
End Function

' // {1A28524C-FDD6-4AA4-AE8F-837EB8267B37}
Public Function CLSID_D2D1ColorManagement() As UUID
    GetMem8 537870295099089.3644@, CLSID_D2D1ColorManagement
    GetMem8 399783166805983.6334@, ByVal VarPtr(CLSID_D2D1ColorManagement) + 8
End Function

' // {921F03D6-641C-47DF-852D-B4BB6153AE11}
Public Function CLSID_D2D1ColorMatrix() As UUID
    GetMem8 517896817037272.7766@, CLSID_D2D1ColorMatrix
    GetMem8 127404742381850.9701@, ByVal VarPtr(CLSID_D2D1ColorMatrix) + 8
End Function

' // {48FC9F51-F6AC-48F1-8B58-3B28AC46F76D}
Public Function CLSID_D2D1Composite() As UUID
    GetMem8 525625345993740.2705@, CLSID_D2D1Composite
    GetMem8 792387977460497.4219@, ByVal VarPtr(CLSID_D2D1Composite) + 8
End Function

' // {407F8C08-5533-4331-A341-23CC3877843E}
Public Function CLSID_D2D1ConvolveMatrix() As UUID
    GetMem8 484174475301378.7656@, CLSID_D2D1ConvolveMatrix
    GetMem8 450485661310407.5171@, ByVal VarPtr(CLSID_D2D1ConvolveMatrix) + 8
End Function

' // {E23F7110-0E9A-4324-AF47-6A2C0C46F35B}
Public Function CLSID_D2D1Crop() As UUID
    GetMem8 483800795808631.6304@, CLSID_D2D1Crop
    GetMem8 662571649489084.6127@, ByVal VarPtr(CLSID_D2D1Crop) + 8
End Function

' // {174319A6-58E9-49B2-BB63-CAF2C811A3DB}
Public Function CLSID_D2D1DirectionalBlur() As UUID
    GetMem8 531040466876413.3798@, CLSID_D2D1DirectionalBlur
    GetMem8 -262023100343501.5237@, ByVal VarPtr(CLSID_D2D1DirectionalBlur) + 8
End Function

' // {90866FCD-488E-454B-AF06-E5041B66C36C}
Public Function CLSID_D2D1DiscreteTransfer() As UUID
    GetMem8 499316438901761.2237@, CLSID_D2D1DiscreteTransfer
    GetMem8 783722004278706.3471@, ByVal VarPtr(CLSID_D2D1DiscreteTransfer) + 8
End Function

' // {EDC48364-0417-4111-9450-43845FA9F890}
Public Function CLSID_D2D1DisplacementMap() As UUID
    GetMem8 468853318788923.4788@, CLSID_D2D1DisplacementMap
    GetMem8 -800045851031769.4828@, ByVal VarPtr(CLSID_D2D1DisplacementMap) + 8
End Function

' // {3E7EFD62-A32D-46D4-A83C-5278889AC954}
Public Function CLSID_D2D1DistantDiffuse() As UUID
    GetMem8 510388369243498.0194@, CLSID_D2D1DistantDiffuse
    GetMem8 610958428042967.364@, ByVal VarPtr(CLSID_D2D1DistantDiffuse) + 8
End Function

' // {428C1EE5-77B8-4450-8AB5-72219C21ABDA}
Public Function CLSID_D2D1DistantSpecular() As UUID
    GetMem8 492256602599011.9141@, CLSID_D2D1DistantSpecular
    GetMem8 -269001939796395.8902@, ByVal VarPtr(CLSID_D2D1DistantSpecular) + 8
End Function

' // {6C26C5C7-34E0-46FC-9CFD-E5823706E228}
Public Function CLSID_D2D1DpiCompensation() As UUID
    GetMem8 511502141527783.9815@, CLSID_D2D1DpiCompensation
    GetMem8 294592394174280.438@, ByVal VarPtr(CLSID_D2D1DpiCompensation) + 8
End Function

' // {61C23C20-AE69-4D8E-94CF-50078DF638F2}
Public Function CLSID_D2D1Flood() As UUID
    GetMem8 558859595524828.2656@, CLSID_D2D1Flood
    GetMem8 -99277263226163.6204@, ByVal VarPtr(CLSID_D2D1Flood) + 8
End Function

' // {409444C4-C419-41A0-B0C1-8CD0C0A18E42}
Public Function CLSID_D2D1GammaTransfer() As UUID
    GetMem8 472899522147570.6052@, CLSID_D2D1GammaTransfer
    GetMem8 479594850270083.5248@, ByVal VarPtr(CLSID_D2D1GammaTransfer) + 8
End Function

' // {1FEB6D69-2FE6-4AC9-8C58-1D7F93E7A6A5}
Public Function CLSID_D2D1GaussianBlur() As UUID
    GetMem8 538889109455001.5337@, CLSID_D2D1GaussianBlur
    GetMem8 -651026159063863.4868@, ByVal VarPtr(CLSID_D2D1GaussianBlur) + 8
End Function

' // {9DAF9369-3846-4D0E-A44E-0C607934A5D7}
Public Function CLSID_D2D1Scale() As UUID
    GetMem8 555243726653879.5881@, CLSID_D2D1Scale
    GetMem8 -290786028849068.0668@, ByVal VarPtr(CLSID_D2D1Scale) + 8
End Function

' // {881DB7D0-F7EE-4D4D-A6D2-4697ACC66EE8}
Public Function CLSID_D2D1Histogram() As UUID
    GetMem8 557038091798509.768@, CLSID_D2D1Histogram
    GetMem8 -169820156489742.2682@, ByVal VarPtr(CLSID_D2D1Histogram) + 8
End Function

' // {0F4458EC-4B32-491B-9E85-BD73F44D3EB6}
Public Function CLSID_D2D1HueRotation() As UUID
    GetMem8 526788686751651.2492@, CLSID_D2D1HueRotation
    GetMem8 -531472479794144.7266@, ByVal VarPtr(CLSID_D2D1HueRotation) + 8
End Function

' // {AD47C8FD-63EF-4ACC-9B51-67979C036C06}
Public Function CLSID_D2D1LinearTransfer() As UUID
    GetMem8 538979273511113.7533@, CLSID_D2D1LinearTransfer
    GetMem8 46274883280223.0683@, ByVal VarPtr(CLSID_D2D1LinearTransfer) + 8
End Function

' // {41251AB7-0BEB-46F8-9DA7-59E93FCCE5DE}
Public Function CLSID_D2D1LuminanceToAlpha() As UUID
    GetMem8 511385048191736.9015@, CLSID_D2D1LuminanceToAlpha
    GetMem8 -238527585275283.6707@, ByVal VarPtr(CLSID_D2D1LuminanceToAlpha) + 8
End Function

' // {EAE6C40D-626A-4C2D-BFCB-391001ABE202}
Public Function CLSID_D2D1Morphology() As UUID
    GetMem8 548915173218155.0093@, CLSID_D2D1Morphology
    GetMem8 20791655386800.4287@, ByVal VarPtr(CLSID_D2D1Morphology) + 8
End Function

' // {6C53006A-4450-4199-AA5B-AD1656FECE5E}
Public Function CLSID_D2D1OpacityMetadata() As UUID
    GetMem8 472688439610749.7578@, CLSID_D2D1OpacityMetadata
    GetMem8 683167733046872.3626@, ByVal VarPtr(CLSID_D2D1OpacityMetadata) + 8
End Function

' // {B9E303C3-C08C-4F91-8B7B-38656BC48C20}
Public Function CLSID_D2D1PointDiffuse() As UUID
    GetMem8 573357551126596.9091@, CLSID_D2D1PointDiffuse
    GetMem8 234546547149193.1019@, ByVal VarPtr(CLSID_D2D1PointDiffuse) + 8
End Function

' // {09C3CA26-3AE2-4F09-9EBC-ED3865D53F22}
Public Function CLSID_D2D1PointSpecular() As UUID
    GetMem8 569514794628754.8966@, CLSID_D2D1PointSpecular
    GetMem8 246792575154583.875@, ByVal VarPtr(CLSID_D2D1PointSpecular) + 8
End Function

' // {06EAB419-DEED-4018-80D2-3E1D471ADEB2}
Public Function CLSID_D2D1Premultiply() As UUID
    GetMem8 461868652747310.3897@, CLSID_D2D1Premultiply
    GetMem8 -555797599739295.68@, ByVal VarPtr(CLSID_D2D1Premultiply) + 8
End Function

' // {5CB2D9CF-327D-459F-A0CE-40C0B2086BF7}
Public Function CLSID_D2D1Saturation() As UUID
    GetMem8 501678402392154.7727@, CLSID_D2D1Saturation
    GetMem8 -61839096001063.7664@, ByVal VarPtr(CLSID_D2D1Saturation) + 8
End Function

' // {C67EA361-1863-4E69-89DB-695D3E9A5B6B}
Public Function CLSID_D2D1Shadow() As UUID
    GetMem8 565007402432401.4945@, CLSID_D2D1Shadow
    GetMem8 773594637758482.7273@, ByVal VarPtr(CLSID_D2D1Shadow) + 8
End Function

' // {818A1105-7932-44F4-AA86-08AE7B2F2C93}
Public Function CLSID_D2D1SpotDiffuse() As UUID
    GetMem8 496872954672513.4597@, CLSID_D2D1SpotDiffuse
    GetMem8 -784184064291159.8934@, ByVal VarPtr(CLSID_D2D1SpotDiffuse) + 8
End Function

' // {EDAE421E-7654-4A37-9DB8-71ACC1BEB3C1}
Public Function CLSID_D2D1SpotSpecular() As UUID
    GetMem8 534787318966270.4158@, CLSID_D2D1SpotSpecular
    GetMem8 -448903466452715.2995@, ByVal VarPtr(CLSID_D2D1SpotSpecular) + 8
End Function

' // {5BF818C3-5E43-48CB-B631-868396D6A1D4}
Public Function CLSID_D2D1TableTransfer() As UUID
    GetMem8 524538983440188.0259@, CLSID_D2D1TableTransfer
    GetMem8 -312498072447836.5258@, ByVal VarPtr(CLSID_D2D1TableTransfer) + 8
End Function

' // {B0784138-3B76-4BC5-B13B-0FA2AD02659F}
Public Function CLSID_D2D1Tile() As UUID
    GetMem8 545983550420944.5176@, CLSID_D2D1Tile
    GetMem8 -696115470425972.8463@, ByVal VarPtr(CLSID_D2D1Tile) + 8
End Function

' // {CF2BB6AE-889A-4AD7-BA29-A2FD732C9FC9}
Public Function CLSID_D2D1Turbulence() As UUID
    GetMem8 539292927728154.795@, CLSID_D2D1Turbulence
    GetMem8 -391836427410091.783@, ByVal VarPtr(CLSID_D2D1Turbulence) + 8
End Function

' // {FB9AC489-AD8D-41ED-9999-BB6347D110F7}
Public Function CLSID_D2D1UnPremultiply() As UUID
    GetMem8 475064400726895.9369@, CLSID_D2D1UnPremultiply
    GetMem8 -64378464216785.8791@, ByVal VarPtr(CLSID_D2D1UnPremultiply) + 8
End Function


