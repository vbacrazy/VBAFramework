VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBA_Path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' .NET Framework System.IO.PathÇÃÉNÉçÅ[Éì
'
' Author: D*isuke YAMAKAWA
' last modified : 2014/1/9

' .NET FrameworkÇ∆éóîÒå›ä∑
Public Function ChangeExtension(path As String, extension As String)
    ChangeExtension = GetFileNameWithoutExtension(path) _
                        & IIf(Left(extension, 1) <> ".", "." & extension, extension)
End Function

Public Function Combine(path1 As String, path2 As String, _
                            Optional path3 As String = "", Optional path4 As String = "")
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Combine = fso.BuildPath(path1, path2)
    Combine = IIf(path3 <> "", fso.BuildPath(Combine, path3), Combine)
    Combine = IIf(path4 <> "", fso.BuildPath(Combine, path4), Combine)
End Function

Public Function GetDirectoryName(path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetDirectoryName = fso.GetParentFolderName(path)
End Function

Public Function GetExtension(path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetExtension = fso.GetExtensionName(path)
End Function

Public Function GetFileName(path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso.GetFileName(path)
End Function

Public Function GetFileNameWithoutExtension(path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameWithoutExtension = fso.GetBaseName(path)
End Function

Public Function GetFullPath(path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetFullPath = fso.GetAbsolutePathName(path)
End Function

Public Function GetTempFileName()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetTempFileName = fso.GetTempName
End Function

Public Function GetTempPath()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetTempPath = fso.GetSpecialFolder(2)
End Function
