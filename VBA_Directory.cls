VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBA_Directory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' .NET Framework System.IO.DirectoryÇÃÉNÉçÅ[Éì
'
' Author: D*isuke YAMAKAWA
' last modified : 2014/1/9

Public Sub CreateDirectory(path As String)
    Call MkDir(path)
End Sub

Public Sub Delete(path As String, Optional recursive As Boolean = False)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.DeleteFolder(path, recursive)
End Sub

' originalÇ…ÇÕë∂ç›ÇπÇ∏
Public Sub Copy(sourcePath As String, destinationPaath As String, recursive As Boolean)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.CopyFolder(sourcePath, destinationPaath, recursive)
End Sub

Public Function Exists(path As String)
    Exists = IIf(Dir(path, vbDirectory) = "", False, True)
End Function

Public Function GetCurrentDirectory()
    GetCurrentDirectory = CurDir
End Function

Public Function GetDirectories(path As String)
    Err.Raise 513, "", "NotImplementedError"
End Function

Public Function GetFiles(path As String, Optional searchPattern As String = "*")
    Err.Raise 513, "", "NotImplementedError"
End Function

Public Sub Move(sourcePath As String, destinationPath As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.moveFolder(sourcePath, destinationPath)
End Sub

Public Sub SetCurrentDirectory(path As String)
    Call ChDrive(path)
    Call ChDir(path)
End Sub
