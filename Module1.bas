Attribute VB_Name = "Module1"
Option Explicit


Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    
Function GetFromInI(KeySection As String, KeyName As String)
    Dim Retstr As String
    Dim IniFileName   As String
    IniFileName = App.Path & "\DIDIDIDIDI.ini"
    Retstr = String(255, vbNullChar)
    GetFromInI = Left(Retstr, GetPrivateProfileString(KeySection, ByVal KeyName, "", Retstr, Len(Retstr), IniFileName))
End Function

Function PutToInI(KeySection As String, KeyName As String, Strings As String)
    Dim IniFileName   As String
    IniFileName = App.Path & "\DIDIDIDIDI.ini"
    Call WritePrivateProfileString(KeySection, KeyName, Strings, IniFileName)
End Function

Public Function DelIniSec(ByVal SectionName As String)      'Çå³ýsection
    Dim IniFileName   As String
    IniFileName = App.Path & "\DIDIDIDIDI.ini"
    Call WritePrivateProfileString(SectionName, 0&, "", IniFileName)
End Function


Public Function GetSectionNames() As String
    Dim IniFileName   As String
    
    Dim sBuf As String
    Dim iLen As Integer
    Dim sLEN As Integer
    sLEN = 2096
    sBuf = String(sLEN, vbNullChar)
    IniFileName = App.Path & "\DIDIDIDIDI.ini"
    
    iLen = GetPrivateProfileSectionNames(sBuf, sLEN, IniFileName)
    
    If iLen <> 0 Then
        If MidB(sBuf, iLen - 1, 1) = vbNullChar Then
            iLen = iLen - 1
        End If
    End If
    
    sBuf = Left(sBuf, iLen)
    
    GetSectionNames = sBuf
'    Dim aSections() As String
'        aSections = Split(sBuf, vbNullChar)
'    'the array aSections now contains section headers

End Function
