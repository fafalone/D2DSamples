VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFontFileEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' // Fonts files enumerator
' // DirectWrite uses that file to enumerate files within fonts collection
' // By The trick

Option Explicit

Implements IDWriteFontFileEnumerator

Private mlCurrentResource   As Long
Private mcFactory           As IDWriteFactory
Private mcLoaders           As Collection

Public Sub Init( _
           ByVal cFactory As IDWriteFactory)
           
    Set mcFactory = cFactory
    
    Set mcLoaders = New Collection
    
    mcLoaders.Add New CFontFileLoader
    mcLoaders.Add New CFontFileLoader
    
    mcFactory.RegisterFontFileLoader mcLoaders(1)
    mcFactory.RegisterFontFileLoader mcLoaders(2)
    
End Sub

Private Sub Class_Initialize()
    
    Debug.Print "CFontFileEnumerator::CFontFileEnumerator"
    mlCurrentResource = -1
    
End Sub

Private Sub Class_Terminate()

    Debug.Print "CFontFileEnumerator::~CFontFileEnumerator"

    mcFactory.UnregisterFontFileLoader mcLoaders(1)
    mcFactory.UnregisterFontFileLoader mcLoaders(2)
    
End Sub

Public Function IDWriteFontFileEnumerator_GetCurrentFontFile() As IDWriteFontFile
    Dim cLoader As CFontFileLoader
    
    ' // Get current file
    If mlCurrentResource >= 0 And mlCurrentResource <= 1 Then
        
        Set cLoader = mcLoaders(mlCurrentResource + 1)
        
        ' // Init file from resource
        cLoader.Init mcFactory, LoadResData(101 + mlCurrentResource, "CUSTOM")
 
        Set IDWriteFontFileEnumerator_GetCurrentFontFile = mcFactory.CreateCustomFontFileReference(0, 0, cLoader)
    Else
    
        Err.Raise 5
        
    End If

End Function

Private Function IDWriteFontFileEnumerator_MoveNext() As Long
    
    ' // Select next file
    mlCurrentResource = mlCurrentResource + 1
    
    If mlCurrentResource > 1 Then
        IDWriteFontFileEnumerator_MoveNext = 0
    Else
        IDWriteFontFileEnumerator_MoveNext = 1
    End If
    
End Function
