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
' last modified : 2014/1/7

Public Function GetFileName(path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileName = fso.GetFileName(path)
End Function

Public Function Combine(path1 As String, path2 As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Combine = fso.BuildPath(path1, path2)
End Function
