Attribute VB_Name = "modMain"
' // Enumerate fonts example

Option Explicit

Private Const LOCALE_NAME_MAX_LENGTH    As Long = 85

Private Declare Function GetUserDefaultLocaleName Lib "kernel32" ( _
                         ByVal lpLocaleName As Long, _
                         ByVal cchLocaleName As Long) As Long

Sub Main()
    Dim cFactory        As IDWriteFactory
    Dim cFontCollection As IDWriteFontCollection
    Dim cFontFamily     As IDWriteFontFamily
    Dim cFamilyNames    As IDWriteLocalizedStrings
    Dim sName           As String
    Dim lIndex          As Long
    Dim lNameIndex      As Long
    Dim lLength         As Long
    Dim sLocaleName     As String
    Dim bFound          As Boolean
    
    sLocaleName = Space$(LOCALE_NAME_MAX_LENGTH)
    lLength = GetUserDefaultLocaleName(StrPtr(sLocaleName), LOCALE_NAME_MAX_LENGTH)
    sLocaleName = Left$(sLocaleName, lLength - 1)
    
    Set cFactory = DW.CreateFactory(DWRITE_FACTORY_TYPE_SHARED)
    
    cFactory.GetSystemFontCollection cFontCollection
    
    For lIndex = 0 To cFontCollection.GetFontFamilyCount - 1
        
        Set cFontFamily = cFontCollection.GetFontFamily(lIndex)
        Set cFamilyNames = cFontFamily.GetFamilyNames

        bFound = cFamilyNames.FindLocaleName(sLocaleName, lNameIndex)
        
        If Not bFound Then
            bFound = cFamilyNames.FindLocaleName("en-us", lNameIndex)
        End If
        
        If Not bFound Then
            lNameIndex = 0
        End If
        
        lLength = cFamilyNames.GetStringLength(lNameIndex)
        
        sName = Space$(lLength)
        
        cFamilyNames.GetString lNameIndex, sName, lLength + 1
        
        ' // Type
        Debug.Print sName

    Next
    
End Sub
