VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFontCollectionLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' // Custom fonts collection loader
' // DirectWrite uses that class to get enumerator of loaded fonts files
' // By The trick

Option Explicit

Implements IDWriteFontCollectionLoaderImp

Private mcEnumerator    As CFontFileEnumerator

Private Sub Class_Initialize()

    Debug.Print "CFontCollectionLoader::CFontCollectionLoader"
    
    Set mcEnumerator = New CFontFileEnumerator
    
End Sub

Private Sub Class_Terminate()
    Debug.Print "CFontCollectionLoader::~CFontCollectionLoader"
End Sub

Private Function IDWriteFontCollectionLoaderImp_CreateEnumeratorFromKey( _
                 ByVal factory As DWVBLib.IDWriteFactory, _
                 ByVal pCollectionKey As Long, _
                 ByVal collectionKeySize As Long) As DWVBLib.IDWriteFontFileEnumerator
    
    ' // Return enumerator of fonts files
    mcEnumerator.Init factory
    
    Set IDWriteFontCollectionLoaderImp_CreateEnumeratorFromKey = mcEnumerator
    
End Function
