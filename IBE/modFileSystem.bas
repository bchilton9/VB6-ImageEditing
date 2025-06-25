Attribute VB_Name = "modFileSystem"
Option Explicit

Public Function GetFileNameOnlyFromFullPath(strFileNameWithFullPath As String) As String
    Dim Parts() As String
    Parts = Split(strFileNameWithFullPath, "\")
    
    GetFileNameOnlyFromFullPath = Parts(UBound(Parts))
End Function

Public Function GetPathOnlyFromFullPath(strFileNameWithFullPath As String) As String
    GetPathOnlyFromFullPath = Left(strFileNameWithFullPath, Len(strFileNameWithFullPath) - Len(GetFileNameOnlyFromFullPath(strFileNameWithFullPath)))
End Function

Public Function CheckPath(PathString As String, Optional AddSlash As Boolean = True) As String
    CheckPath = PathString
    
    If PathString = "" Then
        CheckPath = CheckPath(App.Path, True)
        Exit Function
    End If
    
    If AddSlash And Right(PathString, 1) <> "\" Then CheckPath = PathString & "\"
    If Not AddSlash And Right(PathString, 1) = "\" Then CheckPath = Left(PathString, Len(PathString) - 1)
End Function

