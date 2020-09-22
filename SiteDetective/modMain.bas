Attribute VB_Name = "modMain"
'Another one of my old codes - Sent to planetsourcecode on May 29 2001
'by David Fial - djf1010@aol.com

Option Explicit

Public strSettingFile As String
Public strHeaders As String
Public strServer As String
Public intPort As String
Public intPortProx As String
Public strProxy As String

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Function ReadINI(Section, KeyName, filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function
Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function
Public Sub WriteSet(strKey As String, strValue As String) '
    Call WriteINI("SiteDetective", strKey, strValue, strSettingFile)
End Sub
Public Function ReadSet(strKey As String) As String
    ReadSet = ReadINI("SiteDetective", strKey, strSettingFile)
End Function
