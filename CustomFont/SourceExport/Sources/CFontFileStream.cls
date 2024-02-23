VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFontFileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' // Font file stream object that provides binary data of font
' // By The trick

Option Explicit

Implements IDWriteFontFileStreamImp

Private mbData()    As Byte

Public Sub Init( _
           ByRef bData() As Byte)
    mbData = bData
End Sub

Private Sub Class_Initialize()
    Debug.Print "CFontFileStream::CFontFileStream"
End Sub

Private Sub Class_Terminate()
    Debug.Print "CFontFileStream::~CFontFileStream"
End Sub

Private Function IDWriteFontFileStreamImp_GetFileSize() As LongLong
    IDWriteFontFileStreamImp_GetFileSize = CLngLng(UBound(mbData) + 1) '/ 10000@
End Function

Private Function IDWriteFontFileStreamImp_GetLastWriteTime() As LongLong
    ' // Not implemented
End Function

Public Sub IDWriteFontFileStreamImp_ReadFileFragment( _
            pFragmentStart As LongPtr, _
            ByVal fileOffset As LongLong, _
            ByVal fragmentSize As LongLong, _
            pFragmentContext As LongPtr)
    Dim lOffset As Long
    Dim lSize   As Long

    lOffset = CLng(fileOffset) '* 10000
    lSize = CLng(fragmentSize) '* 10000

    If lOffset + lSize > UBound(mbData) + 1 Then
        Err.Raise E_INVALIDARG
    Else
        pFragmentStart = VarPtr(mbData(lOffset))
    End If

End Sub

Private Sub IDWriteFontFileStreamImp_ReleaseFileFragment( _
            ByVal pFragmentContext As LongPtr)
    ' // Not used
End Sub
