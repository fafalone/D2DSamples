VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFontFileLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' // Font file loader class. That class provides binary stream object with font file data
' // By The trick

Option Explicit

Implements IDWriteFontFileLoaderImp

Private mcFactory   As IDWriteFactory
Private mbData()    As Byte

Public Sub Init( _
           ByVal cFactory As IDWriteFactory, _
           ByRef bData() As Byte)
           
    Set mcFactory = cFactory
    mbData = bData
    
End Sub

Private Sub Class_Initialize()
    Debug.Print "CFontFileLoader::CFontFileLoader"
End Sub

Private Sub Class_Terminate()
    Debug.Print "CFontFileLoader::~CFontFileLoader"
End Sub

Private Function IDWriteFontFileLoaderImp_CreateStreamFromKey( _
                 ByVal pFontFileReferenceKey As Long, _
                 ByVal fontFileReferenceKeySize As Long) As DWVBLib.IDWriteFontFileStream
    Dim cStream As CFontFileStream
    
    Set cStream = New CFontFileStream
    
    cStream.Init mbData
    
    Set IDWriteFontFileLoaderImp_CreateStreamFromKey = cStream
    
End Function
