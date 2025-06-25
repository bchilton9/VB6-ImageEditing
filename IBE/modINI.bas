Attribute VB_Name = "modINI"
Option Explicit

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Function GetINIFile(INIFile As String) As String
If INIFile = "" Then GetINIFile = CheckPath(App.Path, True) & App.EXEName & ".INI" Else GetINIFile = INIFile
End Function

Public Function INIGetLong(Section As String, Key As String, Optional DefaultValue As Long, Optional INIFile As String) As Long
INIGetLong = CLng(GetPrivateProfileInt(Section, Key, DefaultValue, GetINIFile(INIFile)))
End Function

Public Function INIGetString(Section As String, Key As String, Optional DefaultValue As String, Optional INIFile As String) As String
Dim INIBuffer As String
INIBuffer = String(255, " ")
GetPrivateProfileString Section, Key, DefaultValue, INIBuffer, Len(INIBuffer), GetINIFile(INIFile)
INIGetString = Left(Trim(INIBuffer), Len(Trim(INIBuffer)) - 1)
End Function

Public Sub INISet(Section As String, Key As String, Value As Variant, Optional INIFile As String)
WritePrivateProfileString Section, Key, CStr(Value), GetINIFile(INIFile)
End Sub


